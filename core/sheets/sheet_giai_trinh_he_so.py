# -*- coding: utf-8 -*-
"""
Module: core.sheets.sheet_giai_trinh_he_so
Tầng xuất dữ liệu Excel (Presentation Layer) cho Sheet 02_GiaiTrinh_HeSo_DieuChinh (chuẩn F1):
- Giải trình minh bạch từng bước công việc định mức gốc vs bước công việc thực tế áp dụng
- Bảng so sánh 2 cột cạnh nhau (Original vs Applied)
- Công thức động tính tỷ lệ hệ số điều chỉnh k
- Cảnh báo trạng thái vật tư thay thế (Đơn giá = 0 đ)

Tác giả: Tổ Thẩm Định EVN
"""

from typing import Dict, List, Any
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from core.sheets.base import BaseSheetBuilder
from core.constants import SHEET_GIAI_TRINH_HE_SO


class SheetGiaiTrinhHeSoBuilder(BaseSheetBuilder):
    """Builder cho Sheet 02_GiaiTrinh_HeSo_DieuChinh theo chuẩn biểu mẫu F1."""

    def build(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        work_items: List[Dict[str, Any]]
    ) -> None:
        ws = wb.create_sheet(title=SHEET_GIAI_TRINH_HE_SO)

        # 1. Tiêu đề biểu mẫu chuẩn F1
        self.setup_header_banner(
            ws=ws,
            title="BẢNG GIẢI TRÌNH THÀNH PHẦN CÔNG VIỆC VÀ HỆ SỐ ĐIỀU CHỈNH ĐỊNH MỨC",
            project_name=project_name,
            category_name=f"{category_name} (Giải trình chi tiết các biến thể A0, A1, A2... phục vụ thẩm định)",
            max_col=7
        )

        # 2. Nhận diện các biến thể định mức được sử dụng trong dự án
        norm_variants_map = {}
        for it in work_items:
            code = it.get("norm_code", "").strip()
            norm = self.norm_manager.get_norm(code)
            if not norm:
                continue
            eff = self.norm_manager.get_effective_item_resources(it)
            var_id = eff.get("variant_id") or "A0"
            key = (code, var_id)
            if key not in norm_variants_map:
                norm_variants_map[key] = {
                    "code": code,
                    "norm_name": norm.get("name", ""),
                    "norm_unit": norm.get("unit", ""),
                    "legal_basis": norm.get("legal_basis", ""),
                    "variant_id": var_id,
                    "variant_name": eff.get("variant_name") or f"Biến thể {var_id}",
                    "exclude_replacement": eff.get("exclude_replacement", False),
                    "raw_steps": eff.get("raw_steps", []),
                    "k_step": eff.get("k_step", 1.0),
                    "applied_items": []
                }
            norm_variants_map[key]["applied_items"].append(it.get("stt"))

        curr_r = 5
        var_idx = 1
        for (code, var_id), vdata in norm_variants_map.items():
            # Dòng Tiêu Đề Biến Thể
            ws.row_dimensions[curr_r].height = 24
            ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
            t_cell = ws.cell(curr_r, 1, f"MỤC {var_idx}: ĐỊNH MỨC {code} - BIẾN THỂ [{var_id}] ({vdata['variant_name']})")
            t_cell.font = Font(name="Times New Roman", size=11, bold=True, color="002060")
            t_cell.fill = self.fill_header
            t_cell.alignment = self.align_left
            for c in range(1, 8):
                ws.cell(curr_r, c).border = self.border_cell
            curr_r += 1

            # Dòng Phụ: Công tác áp dụng & tình trạng vật tư thay thế
            ws.row_dimensions[curr_r].height = 20
            ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
            items_str = ", ".join(f"STT {s}" for s in vdata["applied_items"] if s)
            vt_status = "ĐÃ LOẠI TRỪ VẬT TƯ THAY THẾ (Đơn giá = 0 đ theo PAKT)" if vdata["exclude_replacement"] else "Có bao gồm vật tư thay thế thiết bị chính"
            sub_cell = ws.cell(curr_r, 1, f"Áp dụng cho các công tác: {items_str} | Trạng thái vật tư thay thế: {vt_status}")
            sub_cell.font = Font(name="Times New Roman", size=10, italic=True)
            sub_cell.alignment = self.align_left
            for c in range(1, 8):
                ws.cell(curr_r, c).border = self.border_cell
            curr_r += 1

            # Tiêu đề bảng so sánh 2 cột
            headers = [
                "Bước", "Nội dung thành phần công việc gốc", "ĐVT",
                "Hao phí gốc (công)", "Áp dụng (1=Có, 0=Không)", "Hao phí áp dụng (công)", "Tỷ lệ áp dụng k"
            ]
            ws.row_dimensions[curr_r].height = 26
            for c_idx, h in enumerate(headers, start=1):
                c_h = ws.cell(curr_r, c_idx, h)
                c_h.font = self.font_header
                c_h.alignment = self.align_center
                c_h.border = self.border_cell
                c_h.fill = self.fill_code if c_idx in [4, 6, 7] else self.fill_header
            curr_r += 1

            start_steps_r = curr_r
            steps = vdata["raw_steps"]
            for idx_s, s in enumerate(steps, start=1):
                ws.row_dimensions[curr_r].height = 20
                st_no = s.get("step") or s.get("step_no") or idx_s
                st_name = s.get("name") or s.get("step_name") or ""
                ws.cell(curr_r, 1, st_no).alignment = self.align_center
                ws.cell(curr_r, 2, st_name).alignment = self.align_left
                ws.cell(curr_r, 3, s.get("dvt_step", "")).alignment = self.align_center

                c_base = ws.cell(curr_r, 4, float(s.get("rate_base", 0.0)))
                c_base.alignment = self.align_right
                c_base.number_format = "0.0000"

                c_flag = ws.cell(curr_r, 5, int(s.get("apply_flag", 1)))
                c_flag.alignment = self.align_center
                c_flag.font = self.font_bold

                c_calc = ws.cell(curr_r, 6, f"=D{curr_r}*E{curr_r}")
                c_calc.alignment = self.align_right
                c_calc.number_format = "0.0000"

                ws.cell(curr_r, 7, "").alignment = self.align_center
                for c in range(1, 8):
                    ws.cell(curr_r, c).border = self.border_cell
                curr_r += 1

            end_steps_r = curr_r - 1

            # Dòng tổng cộng và tỷ lệ k
            ws.row_dimensions[curr_r].height = 22
            ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=3)
            ws.cell(curr_r, 1, "TỔNG CỘNG HAO PHÍ & HỆ SỐ ĐIỀU CHỈNH:").font = self.font_bold
            ws.cell(curr_r, 1).alignment = self.align_right

            c_tot_base = ws.cell(curr_r, 4, f"=SUM(D{start_steps_r}:D{end_steps_r})")
            c_tot_base.font = self.font_bold
            c_tot_base.alignment = self.align_right
            c_tot_base.number_format = "0.0000"

            ws.cell(curr_r, 5, "").alignment = self.align_center

            c_tot_app = ws.cell(curr_r, 6, f"=SUM(F{start_steps_r}:F{end_steps_r})")
            c_tot_app.font = self.font_bold
            c_tot_app.alignment = self.align_right
            c_tot_app.number_format = "0.0000"

            c_k_ratio = ws.cell(curr_r, 7, f"=F{curr_r}/D{curr_r}")
            c_k_ratio.font = Font(name="Times New Roman", size=11, bold=True, color="002060")
            c_k_ratio.fill = self.fill_yellow
            c_k_ratio.alignment = self.align_right
            c_k_ratio.number_format = "0.00%"

            for c in range(1, 8):
                ws.cell(curr_r, c).border = self.border_total
            curr_r += 1

            # Dòng ghi chú nguyên tắc thẩm định
            ws.row_dimensions[curr_r].height = 20
            ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
            c_note = ws.cell(curr_r, 1, "*(Nguyên tắc thẩm định: Tỷ lệ k áp dụng giảm trừ tương ứng cho Vật liệu phụ và Máy thi công; Nhân công tính theo hao phí thực hiện; Vật tư thay thế áp đơn giá = 0 đ theo PAKT)*")
            c_note.font = Font(name="Times New Roman", size=10, italic=True, color="555555")
            c_note.alignment = self.align_left
            for c in range(1, 8):
                ws.cell(curr_r, c).border = self.border_cell
            curr_r += 2
            var_idx += 1

        # Cấu hình Freeze Panes và Page Setup
        self.setup_table_view(ws, freeze_cell="C5", filter_ref=None)
        self.setup_page_printing(ws, orientation="landscape", repeat_rows="4:4")

        col_w = {1: 8, 2: 45, 3: 12, 4: 18, 5: 22, 6: 22, 7: 18}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w
