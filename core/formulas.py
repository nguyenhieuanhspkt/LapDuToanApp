# -*- coding: utf-8 -*-
"""
Module: core.formulas
Tập trung toàn bộ việc sinh công thức Excel động (Dynamic Formulas) cho dự án.

Đảm bảo:
- Không viết chuỗi công thức phân tán rải rác trong mã nguồn.
- Đồng nhất cú pháp tham chiếu giữa các Sheet trong Excel.
- Dễ dàng thay đổi cấu trúc biểu mẫu hoặc kiểm thử độc lập.

Tác giả: Tổ Thẩm Định EVN
"""

from typing import List, Union
from core.constants import LET_COL_AMOUNT, LET_COL_RATE, LET_COL_PRICE, LET_COL_EVIDENCE_RATE


class ExcelFormulaBuilder:
    """Lớp tiện ích chuyên trách xây dựng các chuỗi công thức Excel chuẩn xác."""

    @staticmethod
    def multiply_cells(col_a: str, col_b: str, row: int) -> str:
        """Sinh công thức nhân hai ô: =A1*B1."""
        return f"={col_a}{row}*{col_b}{row}"

    @staticmethod
    def item_amount(row: int) -> str:
        """Sinh công thức thành tiền công tác / tài nguyên: =G{row}*H{row}."""
        return f"={LET_COL_RATE}{row}*{LET_COL_PRICE}{row}"

    @staticmethod
    def sum_range(col: str, start_row: int, end_row: int) -> str:
        """Sinh công thức hàm tổng một dải ô: =SUM(J6:J12)."""
        return f"=SUM({col}{start_row}:{col}{end_row})"

    @staticmethod
    def sum_discrete_cells(cells: List[str]) -> str:
        """Sinh công thức cộng các ô rời rạc: =J10+J20+J30."""
        if not cells:
            return "=0"
        return "=" + "+".join(cells)

    @staticmethod
    def evidence_rate_link(
        norm_code: str,
        evidence_row: int,
        divisor: float = 1.0,
        k_step: float = 1.0
    ) -> str:
        """
        Sinh công thức liên kết định mức hao phí đến Sheet Minh chứng tra cứu.
        
        Ví dụ:
        - Chuẩn gốc: ='BD.41111'!$R$15
        - Có biến thể: ='BD.41111'!$R$15*0.6
        - Có quy đổi đơn vị: ='SB.822'!$R$12/10.0
        """
        base_ref = f"='{norm_code}'!${LET_COL_EVIDENCE_RATE}${evidence_row}"
        
        formula = base_ref
        if k_step < 1.0:
            formula += f"*{k_step}"
        if divisor != 1.0:
            formula += f"/{divisor}"
            
        return formula

    @staticmethod
    def th_price_link(th_sheet_name: str, price_col: str, th_row: int) -> str:
        """
        Sinh công thức liên kết đơn giá từ các bảng Tổng hợp tài nguyên (TH_VatLieu, TH_NhanCong, TH_MayThiCong).
        
        Ví dụ: ='TH_VatLieu'!$G$12 hoặc ='TH_NhanCong'!$F$8
        """
        return f"='{th_sheet_name}'!${price_col}${th_row}"

    @staticmethod
    def percentage_material_amount(
        amount_col: str,
        mat_start_row: int,
        prev_row: int,
        rate_col: str,
        curr_row: int
    ) -> str:
        """
        Sinh công thức tính thành tiền vật liệu phụ theo tỷ lệ % trên tổng vật liệu chính:
        Ví dụ: =SUM(J7:J12)*G13/100
        """
        return f"=SUM({amount_col}{mat_start_row}:{amount_col}{prev_row})*{rate_col}{curr_row}/100"
