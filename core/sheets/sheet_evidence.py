"""
Sheet Builder: Individual Evidence Sheets (BD.41111, E6C.3.05, etc.)
Manages embedded scan images at B3, work step breakdown, and unit rate tables.
Author: Tổ Thẩm Định EVN
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import Dict, List, Any, Optional
from core.sheets.base import BaseSheetBuilder
from openpyxl.drawing.image import Image as OpenpyxlImage
from config import IMAGES_DIR
from core.norm_manager import REPLACEMENT_MATERIAL_CODES

class SheetEvidenceBuilder(BaseSheetBuilder):
    def build(self, wb: openpyxl.Workbook, code: str) -> Dict[str, Any]:
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

        # Setup page printing for evidence audit
        self.setup_page_printing(ws, orientation="landscape", repeat_rows="3:3")

        return cell_coords


