"""
Excel Generator Module for LapDuToanApp.
Generates complete Excel workbooks with 100% dynamic formula links:
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
        output_path: Path
    ) -> Path:
        """
        Main generator method: builds the complete workbook and writes to output_path.
        """
        wb = openpyxl.Workbook()
        # Remove default sheet
        wb.remove(wb.active)

        # 1. Determine all unique norm codes used in the work items
        unique_norm_codes = []
        for item in work_items:
            code = item["norm_code"]
            if code not in unique_norm_codes and code in self.norm_manager.norms:
                unique_norm_codes.append(code)

        # 2. Build 3 Resource Lookup Tables (for TH_VatLieu, TH_NhanCong, TH_MayThiCong)
        res_data = self.norm_manager.aggregate_resources_from_norms()
        # Filter to only resources actually used in the selected work items
        used_mat_codes = set()
        used_lab_codes = set()
        used_mac_codes = set()

        for code in unique_norm_codes:
            norm = self.norm_manager.get_norm(code)
            res = norm.get("resources", {})
            for m in res.get("materials", []):
                if m.get("code"): used_mat_codes.add(m["code"])
            for l in res.get("labor", []):
                if l.get("code"): used_lab_codes.add(l["code"])
            for mac in res.get("machines", []):
                if mac.get("code"): used_mac_codes.add(mac["code"])

        # Retain order from database
        materials_list = [m for m in res_data["materials"] if m["code"] in used_mat_codes or not used_mat_codes]
        labor_list = [l for l in res_data["labor"] if l["code"] in used_lab_codes or not used_lab_codes]
        machines_list = [mac for mac in res_data["machines"] if mac["code"] in used_mac_codes or not used_mac_codes]

        # Row mappings in TH sheets (headers start at row 6, data rows start at row 7)
        mat_row_map = {m["code"]: 7 + idx for idx, m in enumerate(materials_list)}
        lab_row_map = {l["code"]: 7 + idx for idx, l in enumerate(labor_list)}
        mac_row_map = {mac["code"]: 7 + idx for idx, mac in enumerate(machines_list)}

        # 3. Create Sheet: 00_ThongKe_DinhMuc
        self._build_sheet_thong_ke(wb, unique_norm_codes)

        # 4. Create Evidence Sheets for each unique norm code
        # Keep track of cell coordinates where each norm's resources are located
        norm_cell_map = {}
        for code in unique_norm_codes:
            norm_cell_map[code] = self._build_evidence_sheet(wb, code)

        # 5. Create Sheet: 01_ChiTiet_HaoPhi_DinhMuc
        self._build_sheet_don_gia_chi_tiet(
            wb, project_name, category_name, work_items,
            norm_cell_map, mat_row_map, lab_row_map, mac_row_map
        )

        # 6. Create Sheet: TH_VatLieu
        self._build_sheet_th_vat_lieu(wb, project_name, category_name, materials_list)

        # 7. Create Sheet: TH_NhanCong
        self._build_sheet_th_nhan_cong(wb, project_name, category_name, labor_list)

        # 8. Create Sheet: TH_MayThiCong
        self._build_sheet_th_may_thi_cong(wb, project_name, category_name, machines_list)

        # Reorder sheets cleanly: 00_ThongKe -> 01_ChiTiet -> TH_VL -> TH_NC -> TH_M -> Evidence sheets
        ordered_names = ["00_ThongKe_DinhMuc", "01_ChiTiet_HaoPhi_DinhMuc", "TH_VatLieu", "TH_NhanCong", "TH_MayThiCong"]
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
                # Scale appropriately
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
        ws["P3"] = "Thành phần công việc"
        ws["Q3"] = "ĐVT / Tỷ lệ"
        ws["R3"] = "Định mức gốc"
        ws["S3"] = "Áp dụng (1=Có, 0=Không)"
        ws["T3"] = "Hao phí tính toán"
        for col_idx, col_let in enumerate(["P", "Q", "R", "S", "T"], start=16):
            cell = ws[f"{col_let}3"]
            cell.font = self.font_ev_header
            cell.fill = self.fill_ev_header
            cell.border = self.border_cell
            cell.alignment = self.align_center

        steps = norm.get("work_steps", [])
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
            # Formula =R*S
            c_t = ws.cell(r_step, 20, f"=R{r_step}*S{r_step}")
            c_t.alignment = self.align_right
            c_t.font = self.font_ev_reg
            for c in range(16, 21):
                ws.cell(r_step, c).border = self.border_cell
            r_step += 1

        # Total step row
        r_tot_step = r_step
        ws.cell(r_tot_step, 16, "Tổng cộng hao phí các bước:").font = self.font_ev_bold
        ws.cell(r_tot_step, 16).alignment = self.align_right
        ws.cell(r_tot_step, 18, f"=SUM(R4:R{r_tot_step-1})").font = self.font_ev_bold
        ws.cell(r_tot_step, 20, f"=SUM(T4:T{r_tot_step-1})").font = self.font_ev_bold
        for c in range(16, 21):
            ws.cell(r_tot_step, c).border = self.border_cell

        # 4. Resource Table (Vật liệu, Nhân công, Máy thi công)
        r_res_title = r_tot_step + 2
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
        mats = res.get("materials", [])
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
        labs = res.get("labor", [])
        if labs:
            cat_lbl = "2. Nhân công" if mats else "1. Nhân công trực tiếp"
            ws.cell(curr_r, 16, cat_lbl).font = self.font_ev_cat
            for c in range(16, 21): ws.cell(curr_r, c).border = self.border_cell
            curr_r += 1
            for l in labs:
                ws.cell(curr_r, 16, f"   - {l['name']} ({l['code']})").font = self.font_ev_reg
                ws.cell(curr_r, 17, l.get("unit", "công")).alignment = self.align_center
                # Link to total step if formula_ref or SCL
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
        macs = res.get("machines", [])
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

        # Set column widths
        col_w = {"P": 45, "Q": 12, "R": 18, "S": 22, "T": 25}
        for col_l, w in col_w.items():
            ws.column_dimensions[col_l].width = w

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
        for item in work_items:
            stt = item["stt"]
            code = item["norm_code"]
            norm = self.norm_manager.get_norm(code)
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
            mats = res.get("materials", [])
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
                    
                    # Col G: Link to evidence sheet
                    r_ev = coords["materials"].get(m_code)
                    if r_ev:
                        ws.cell(curr_r, 7, f"='{code}'!R{r_ev}").alignment = self.align_right
                    else:
                        ws.cell(curr_r, 7, m.get("rate", 0.0)).alignment = self.align_right
                    ws.cell(curr_r, 7).number_format = "#,##0.00"

                    # Col H: Link to TH_VatLieu
                    if m_code == "V00750" or m.get("unit") == "%":
                        # Percent material, no direct price
                        ws.cell(curr_r, 10, f"=SUM(J{mat_start_r}:J{curr_r-1})*G{curr_r}/100").alignment = self.align_right
                    else:
                        r_th = mat_row_map.get(m_code)
                        if r_th:
                            c_h = ws.cell(curr_r, 8, f"='TH_VatLieu'!$G${r_th}")
                        else:
                            c_h = ws.cell(curr_r, 8, 0)
                        c_h.fill = self.fill_yellow
                        c_h.alignment = self.align_right
                        c_h.number_format = "#,##0"
                        ws.cell(curr_r, 10, f"=G{curr_r}*H{curr_r}").alignment = self.align_right

                    ws.cell(curr_r, 10).number_format = "#,##0"
                    curr_r += 1

                mat_end_r = curr_r - 1
                # Subtotal for materials in Job Row
                ws.cell(job_hdr_r + 1, 10, f"=SUM(J{mat_start_r}:J{mat_end_r})").font = self.font_cat
                ws.cell(job_hdr_r + 1, 10).number_format = "#,##0"
                subtotal_rows.append(job_hdr_r + 1)

            # 2. Labor
            labs = res.get("labor", [])
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
                    
                    # Col G: Link to evidence sheet
                    r_ev = coords["labor"].get(l_code)
                    if r_ev:
                        ws.cell(curr_r, 7, f"='{code}'!R{r_ev}").alignment = self.align_right
                    else:
                        ws.cell(curr_r, 7, l.get("rate", 0.0)).alignment = self.align_right
                    ws.cell(curr_r, 7).number_format = "#,##0.00"

                    # Col H: Link to TH_NhanCong
                    r_th = lab_row_map.get(l_code)
                    if r_th:
                        c_h = ws.cell(curr_r, 8, f"='TH_NhanCong'!$G${r_th}")
                    else:
                        c_h = ws.cell(curr_r, 8, 0)
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
            macs = res.get("machines", [])
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
                    
                    # Col G: Link to evidence sheet
                    r_ev = coords["machines"].get(m_code)
                    if r_ev:
                        ws.cell(curr_r, 7, f"='{code}'!R{r_ev}").alignment = self.align_right
                    else:
                        ws.cell(curr_r, 7, mac.get("rate", 0.0)).alignment = self.align_right
                    ws.cell(curr_r, 7).number_format = "#,##0.00"

                    # Col H: Link to TH_MayThiCong
                    r_th = mac_row_map.get(m_code)
                    if r_th:
                        c_h = ws.cell(curr_r, 8, f"='TH_MayThiCong'!$I${r_th}")
                    else:
                        c_h = ws.cell(curr_r, 8, 0)
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

            # Overall unit rate for this work item in job header
            if subtotal_rows:
                ws.cell(job_hdr_r, 10, "=" + "+".join(f"J{r_sub}" for r_sub in subtotal_rows)).font = self.font_bold
                ws.cell(job_hdr_r, 10).number_format = "#,##0"

        # Apply borders to entire table area
        for r in range(5, curr_r):
            for c in range(1, 11):
                ws.cell(r, c).border = self.border_cell

        col_w = {1: 6, 2: 14, 3: 4, 4: 4, 5: 50, 6: 12, 7: 15, 8: 15, 9: 10, 10: 18}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _build_sheet_th_vat_lieu(self, wb: openpyxl.Workbook, project_name: str, category_name: str, materials: List[Dict[str, Any]]):
        ws = wb.create_sheet(title="TH_VatLieu")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG TỔNG HỢP VẬT LIỆU VÀ CHÊNH LỆCH GIÁ"
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

        headers = ["STT", "Mã hiệu", "Tên vật tư", "Đơn vị", "Khối lượng", "Giá gốc", "Giá thông báo", "Chênh lệch", "Tổng chênh"]
        ws.row_dimensions[6].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(6, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            if c_idx == 2: cell.fill = self.fill_code

        r_start = 7
        for idx, m in enumerate(materials, start=1):
            r = r_start + idx - 1
            ws.row_dimensions[r].height = 20
            ws.cell(r, 1, idx).alignment = self.align_center
            
            c_code = ws.cell(r, 2, m["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code
            
            ws.cell(r, 3, m["name"]).alignment = self.align_left
            ws.cell(r, 4, m["unit"]).alignment = self.align_center
            
            c_kl = ws.cell(r, 5, 0)
            c_kl.alignment = self.align_right
            c_kl.number_format = "#,##0.00"
            
            c_goc = ws.cell(r, 6, m.get("price_base", 0))
            c_goc.alignment = self.align_right
            c_goc.number_format = "#,##0"
            
            c_tb = ws.cell(r, 7, m.get("price_notice", 0))
            c_tb.alignment = self.align_right
            c_tb.number_format = "#,##0"
            
            c_cl = ws.cell(r, 8, f"=G{r}-F{r}")
            c_cl.alignment = self.align_right
            c_cl.number_format = "#,##0"
            
            c_tc = ws.cell(r, 9, f"=E{r}*H{r}")
            c_tc.alignment = self.align_right
            c_tc.number_format = "#,##0"
            
            for c in range(1, 10):
                ws.cell(r, c).border = self.border_cell
                if c not in [2]: ws.cell(r, c).font = self.font_regular

        r_total = r_start + len(materials)
        ws.row_dimensions[r_total].height = 22
        for c in range(1, 10): ws.cell(r_total, c).border = self.border_total
        ws.cell(r_total, 3, "Tổng vật liệu:").font = self.font_total
        ws.cell(r_total, 3).alignment = Alignment(horizontal="right", vertical="center")
        c_tot = ws.cell(r_total, 9, f"=SUM(I{r_start}:I{r_total-1})")
        c_tot.font = self.font_total
        c_tot.alignment = self.align_right
        c_tot.number_format = "#,##0"

        col_w = {1: 6, 2: 14, 3: 45, 4: 10, 5: 14, 6: 15, 7: 15, 8: 15, 9: 18}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _build_sheet_th_nhan_cong(self, wb: openpyxl.Workbook, project_name: str, category_name: str, labor: List[Dict[str, Any]]):
        ws = wb.create_sheet(title="TH_NhanCong")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG TỔNG HỢP NHÂN CÔNG VÀ CHÊNH LỆCH GIÁ"
        ws["A1"].font = self.font_title
        ws["A2"] = "(Căn cứ theo 5898/SXD-QLXD ngày 02/7/2026 của SXD tỉnh Lâm Đồng)"
        ws["A2"].font = self.font_subtitle
        ws["A3"] = f"CÔNG TRÌNH: {project_name}"
        ws["A3"].font = self.font_project
        ws["A4"] = f"HẠNG MỤC: {category_name}"
        ws["A4"].font = self.font_project
        ws["I5"] = "Đơn vị: đồng"
        ws["I5"].font = self.font_unit
        ws["I5"].alignment = self.align_right

        headers = ["STT", "Mã hiệu", "Tên vật tư", "Đơn vị", "Khối lượng", "Giá gốc", "Giá hiện tại", "Chênh lệch", "Tổng chênh"]
        ws.row_dimensions[6].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(6, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            if c_idx == 2: cell.fill = self.fill_code

        r_start = 7
        for idx, l in enumerate(labor, start=1):
            r = r_start + idx - 1
            ws.row_dimensions[r].height = 20
            ws.cell(r, 1, idx).alignment = self.align_center
            
            c_code = ws.cell(r, 2, l["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code
            
            ws.cell(r, 3, l["name"]).alignment = self.align_left
            ws.cell(r, 4, l["unit"]).alignment = self.align_center
            
            c_kl = ws.cell(r, 5, 0)
            c_kl.alignment = self.align_right
            c_kl.number_format = "#,##0.00"
            
            c_goc = ws.cell(r, 6, l.get("price_base", 0))
            c_goc.alignment = self.align_right
            c_goc.number_format = "#,##0"
            
            c_ht = ws.cell(r, 7, l.get("price_current", 0))
            c_ht.alignment = self.align_right
            c_ht.font = self.font_blue
            c_ht.number_format = "#,##0"
            
            c_cl = ws.cell(r, 8, f"=G{r}-F{r}")
            c_cl.alignment = self.align_right
            c_cl.number_format = "#,##0"
            
            c_tc = ws.cell(r, 9, f"=E{r}*H{r}")
            c_tc.alignment = self.align_right
            c_tc.number_format = "#,##0"
            
            for c in range(1, 10):
                ws.cell(r, c).border = self.border_cell
                if c not in [2, 7]: ws.cell(r, c).font = self.font_regular

        r_total = r_start + len(labor)
        ws.row_dimensions[r_total].height = 22
        for c in range(1, 10): ws.cell(r_total, c).border = self.border_total
        ws.cell(r_total, 3, "Cộng nhân công:").font = self.font_total
        ws.cell(r_total, 3).alignment = Alignment(horizontal="right", vertical="center")
        c_tot = ws.cell(r_total, 9, f"=SUM(I{r_start}:I{r_total-1})")
        c_tot.font = self.font_total
        c_tot.alignment = self.align_right
        c_tot.number_format = "#,##0"

        col_w = {1: 6, 2: 14, 3: 45, 4: 10, 5: 14, 6: 15, 7: 15, 8: 15, 9: 18}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _build_sheet_th_may_thi_cong(self, wb: openpyxl.Workbook, project_name: str, category_name: str, machines: List[Dict[str, Any]]):
        ws = wb.create_sheet(title="TH_MayThiCong")
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "BẢNG TỔNG HỢP MÁY THI CÔNG VÀ CHÊNH LỆCH GIÁ"
        ws["A1"].font = self.font_title
        ws["A2"] = "(Căn cứ theo TT 37/2026/TT-BXD ngày 26/06/2026 của Bộ Xây dựng)"
        ws["A2"].font = self.font_subtitle
        ws["A3"] = f"CÔNG TRÌNH: {project_name}"
        ws["A3"].font = self.font_project
        ws["A4"] = f"HẠNG MỤC: {category_name}"
        ws["A4"].font = self.font_project
        ws["K5"] = "Đơn vị: đồng"
        ws["K5"].font = self.font_unit
        ws["K5"].alignment = self.align_right

        headers = ["STT", "Mã hiệu", "Tên vật tư", "Đơn vị", "Khối lượng", "Giá gốc", "Giá thông báo", "Bù giá", "Giá hiện tại", "Chênh lệch", "Tổng chênh"]
        ws.row_dimensions[6].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(6, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            if c_idx == 2: cell.fill = self.fill_code

        r_start = 7
        for idx, mac in enumerate(machines, start=1):
            r = r_start + idx - 1
            ws.row_dimensions[r].height = 20
            ws.cell(r, 1, idx).alignment = self.align_center
            
            c_code = ws.cell(r, 2, mac["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code
            
            ws.cell(r, 3, mac["name"]).alignment = self.align_left
            ws.cell(r, 4, mac["unit"]).alignment = self.align_center
            
            c_kl = ws.cell(r, 5, 0)
            c_kl.alignment = self.align_right
            c_kl.number_format = "#,##0.00"
            
            c_goc = ws.cell(r, 6, mac.get("price_base", 0))
            c_goc.alignment = self.align_right
            c_goc.number_format = "#,##0"
            
            c_tb = ws.cell(r, 7, mac.get("price_notice", 0))
            c_tb.alignment = self.align_right
            c_tb.number_format = "#,##0"
            
            c_bu = ws.cell(r, 8, mac.get("price_diff", 0))
            c_bu.alignment = self.align_right
            c_bu.number_format = "#,##0"
            
            c_ht = ws.cell(r, 9, f"=G{r}+H{r}")
            c_ht.alignment = self.align_right
            c_ht.font = self.font_blue
            c_ht.number_format = "#,##0"
            
            c_cl = ws.cell(r, 10, f"=I{r}-F{r}")
            c_cl.alignment = self.align_right
            c_cl.number_format = "#,##0"
            
            c_tc = ws.cell(r, 11, f"=E{r}*J{r}")
            c_tc.alignment = self.align_right
            c_tc.number_format = "#,##0"
            
            for c in range(1, 12):
                ws.cell(r, c).border = self.border_cell
                if c not in [2, 9]: ws.cell(r, c).font = self.font_regular

        r_total = r_start + len(machines)
        ws.row_dimensions[r_total].height = 22
        for c in range(1, 12): ws.cell(r_total, c).border = self.border_total
        ws.cell(r_total, 3, "Cộng Máy:").font = self.font_total
        ws.cell(r_total, 3).alignment = Alignment(horizontal="right", vertical="center")
        c_tot = ws.cell(r_total, 11, f"=SUM(K{r_start}:K{r_total-1})")
        c_tot.font = self.font_total
        c_tot.alignment = self.align_right
        c_tot.number_format = "#,##0"

        col_w = {1: 6, 2: 14, 3: 48, 4: 10, 5: 14, 6: 15, 7: 15, 8: 12, 9: 15, 10: 15, 11: 18}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w
