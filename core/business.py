# -*- coding: utf-8 -*-
"""
Module: core.business
Tầng nghiệp vụ (Business / Domain Layer) của hệ thống thẩm định dự toán.

Đảm nhận:
- Toàn bộ các hàm kiểm tra logic nghiệp vụ (Business Predicates)
- Quy tắc điều chỉnh định mức, biến thể công việc (k_step, variant)
- Quy tắc vật tư thay thế (exclude_replacement)
- Quy tắc vật tư phụ tính theo tỷ lệ phần trăm (V00750, %)
- Quy đổi đơn vị tính (divisor giữa định mức gốc và khối lượng nghiệm thu)

Tác giả: Tổ Thẩm Định EVN
"""

from typing import Dict, Any, Optional
from core.constants import (
    PERCENTAGE_MATERIAL_CODE,
    REPLACEMENT_MATERIAL_CODES,
    QUOTE_CODE_PREFIXES
)


def is_quote_item(item: Dict[str, Any], norm: Optional[Dict[str, Any]]) -> bool:
    """
    Xác định công tác có phải là công tác báo giá / tạm tính hay không.
    
    Quy tắc nghiệp vụ:
    - Nếu công tác được đánh dấu tường minh is_quote=True.
    - Hoặc không tra cứu được mã hiệu trong cơ sở dữ liệu định mức.
    - Hoặc mã hiệu bắt đầu bằng các tiền tố báo giá/tạm tính: BG (Báo giá), TT (Tạm tính), VT (Vật tư).
    """
    if item.get("is_quote", False):
        return True
    if not norm:
        return True
    code = (item.get("norm_code") or "").strip().upper()
    return any(code.startswith(prefix) for prefix in QUOTE_CODE_PREFIXES)


def is_replacement_material(code: str, exclude_replacement: bool) -> bool:
    """
    Kiểm tra vật tư có thuộc diện vật tư thay thế cần loại trừ đơn giá hay không.
    
    Quy tắc nghiệp vụ:
    - Trong công tác bảo dưỡng định kỳ, đơn vị thi công chỉ cung cấp nhân công dịch vụ.
    - Các vật tư thay thế chính (như bình ắc quy, bo mạch, cảm biến...) do Chủ đầu tư
      cấp hoặc mua sắm theo gói thiết bị riêng, do đó trong dự toán dịch vụ phải đặt Đơn giá = 0.
    - Các mã thuộc REPLACEMENT_MATERIAL_CODES khi có cờ exclude_replacement sẽ được loại trừ.
    """
    return bool(exclude_replacement and code in REPLACEMENT_MATERIAL_CODES)


def is_percentage_material(mat: Dict[str, Any]) -> bool:
    """
    Kiểm tra vật tư có phải là 'Vật liệu phụ' tính theo tỷ lệ % hay không.
    
    Quy tắc nghiệp vụ:
    - Định mức Bộ Xây dựng quy định vật liệu phụ (như giẻ lau, cồn tẩy, băng keo...)
      được tính bằng tỷ lệ % (thường là 3% - 5%) trên tổng chi phí các vật liệu chính.
    - Nhận diện qua mã V00750 hoặc đơn vị tính là '%'.
    - Thành tiền trong Excel được tính bằng công thức: =SUM(J_start:J_end) * G_rate / 100.
    """
    code = (mat.get("code") or "").strip()
    unit = (mat.get("unit") or "").strip()
    return code == PERCENTAGE_MATERIAL_CODE or unit == "%"


def get_unit_divisor(norm_unit: str, item_unit: str) -> float:
    """
    Tính hệ số quy đổi đơn vị (Divisor) giữa đơn vị định mức gốc và đơn vị công tác.
    
    Quy tắc nghiệp vụ:
    - Nhiều định mức Nhà nước quy định hao phí cho 10 đơn vị (ví dụ '10 đầu báo', '10 bình', '100m').
    - Nếu khối lượng công tác dự án tính theo 1 đơn vị ('đầu', 'cái', 'bình'):
      Hao phí cho 1 đơn vị = Hao phí định mức gốc / 10 (divisor = 10.0).
    - Nếu khối lượng công tác dự án cũng tính theo '10 đầu': divisor = 1.0 (không chia).
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


def should_link_evidence_sheet(
    has_evidence_row: bool,
    is_replacement: bool,
    is_percentage: bool
) -> bool:
    """
    Xác định ô định mức hao phí tại Cột G có nên link động đến Sheet Minh chứng hay không.
    
    Quy tắc:
    - Phải có dòng tọa độ tương ứng trên Sheet Minh chứng (has_evidence_row is True).
    - Không áp dụng link cho vật tư thay thế bị loại trừ (đơn giá = 0).
    - Không áp dụng link trực tiếp cho vật tư phụ dạng % (vì tỷ lệ % tính trên tổng).
    """
    return bool(has_evidence_row and not is_replacement and not is_percentage)


def calculate_effective_rate(base_rate: float, divisor: float = 1.0, k_step: float = 1.0) -> float:
    """
    Tính trị số định mức hiệu dụng (dùng cho giá trị số tĩnh hoặc kiểm tra kiểm toán).
    
    Công thức: Effective_Rate = (Base_Rate * k_step) / Divisor
    """
    div = divisor if divisor != 0 else 1.0
    return float(base_rate) * float(k_step) / float(div)


def format_job_display_name(
    item: Dict[str, Any],
    norm: Optional[Dict[str, Any]],
    eff: Dict[str, Any]
) -> str:
    """
    Chuẩn hóa tên hiển thị của công tác kèm theo chú thích biến thể và loại trừ vật tư.
    
    Ví dụ: 'Bảo dưỡng trung tâm báo cháy [Biến thể A1: 60% thành phần ĐM] (Không tính vật tư thay thế)'
    """
    job_name = item.get("job_name", norm.get("name", "") if norm else "")
    k_step = eff.get("k_step", 1.0)
    var_id = item.get("variant_id", eff.get("variant_id", "A0"))

    if k_step < 1.0:
        job_name += f" [Biến thể {var_id}: {round(k_step * 100)}% thành phần ĐM]"
    if eff.get("exclude_replacement"):
        job_name += " (Không tính vật tư thay thế)"

    return job_name
