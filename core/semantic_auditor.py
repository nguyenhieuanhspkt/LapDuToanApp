"""
Core Semantic Auditor & Raw Data Analyzer for PAKT
Module Tiền Thẩm Tra & Bóc Tách Bản Chất Kỹ Thuật Dữ Liệu Thô PAKT (Độc Lập)
Author: Tổ Thẩm Định EVN
"""

import re
from typing import List, Dict, Any, Tuple
from collections import Counter
from core.norm_manager import NormManager

class SemanticAuditor:
    """
    Bộ bóc tách ngữ nghĩa V-O-S-C, suy luận bản chất thi công,
    chấm điểm độ rõ ràng, cảnh báo 'Cần làm rõ gì?' và đối chiếu nội hàm định mức.
    """

    def __init__(self, norm_manager: NormManager = None):
        self.norm_manager = norm_manager or NormManager()
        self._init_taxonomies()

    def _init_taxonomies(self):
        # 1. Action Verbs (Hành động V)
        self.verb_patterns = [
            ("Lắp mới / Thay mới", [
                r"\blắp mới\b", r"\bthay mới\b", r"\blắp đặt mới\b", r"\bcung cấp và lắp đặt\b", 
                r"\bmua mới\b", r"\btrang bị mới\b"
            ]),
            ("Thay thế linh kiện", [
                r"\bthay thế\b", r"\bthay\b", r"\bthay pin\b", r"\bthay ắc quy\b", r"\bđổi\b"
            ]),
            ("Tháo dỡ / Thu hồi", [
                r"\btháo dỡ\b", r"\btháo\b", r"\bthu hồi\b", r"\btháo bỏ\b", r"\bgỡ\b"
            ]),
            ("Bảo dưỡng / Sửa chữa", [
                r"\bbảo dưỡng\b", r"\bbảo trì\b", r"\bsửa chữa\b", r"\btra mỡ\b", r"\bvệ sinh\b", 
                r"\blàm sạch\b", r"\bsúc rửa\b", r"\bthổi bụi\b", r"\bxịt rp7\b", r"\bvô mỡ\b"
            ]),
            ("Đo kiểm / Thí nghiệm", [
                r"\bthí nghiệm\b", r"\bthử nghiệm\b", r"\bđo kiểm\b", r"\bkiểm tra\b", r"\bđo nội trở\b",
                r"\bđo điện trở\b", r"\bthông mạch\b", r"\bgiả lập tác động\b", r"\bthử tác động\b"
            ]),
            ("Hiệu chỉnh / Căn chỉnh", [
                r"\bhiệu chỉnh\b", r"\bhiệu chuẩn\b", r"\bđộ nhạy\b", r"\bcăn chỉnh\b", r"\bcân chỉnh\b",
                r"\bbơm khí chuẩn\b", r"\bzero span\b"
            ])
        ]

        # 2. Target Objects (Đối tượng kỹ thuật O)
        self.object_patterns = [
            ("Đầu báo khói", [r"\bđầu báo khói\b", r"\bbuồng khói\b", r"\bbáo khói\b", r"\bsmoke detector\b"]),
            ("Đầu báo nhiệt", [r"\bđầu báo nhiệt\b", r"\bbáo nhiệt\b", r"\bheat detector\b", r"\bgia tăng\b"]),
            ("Đầu báo khói/nhiệt hỗn hợp", [r"\bđầu báo khói nhiệt\b", r"\bđầu báo và đầu báo cháy\b", r"\bđầu báo\b"]),
            ("Nút nhấn khẩn báo cháy", [r"\bnút nhấn khẩn\b", r"\bnút ấn khẩn\b", r"\bnút nhấn báo cháy\b", r"\bnút nhấn\b", r"\bnút ấn\b", r"\bmanual call point\b"]),
            ("Đèn báo cháy", [r"\bđèn báo cháy\b", r"\bđèn báo\b", r"\bđèn vị trí\b", r"\bindicator\b"]),
            ("Chuông / Còi báo cháy", [r"\bchuông báo cháy\b", r"\bchuông\b", r"\bcòi báo cháy\b", r"\bcòi\b", r"\bhorn\b", r"\bbell\b"]),
            ("Module giám sát / Điều khiển", [r"\bmodule điều khiển\b", r"\bmodule giám sát\b", r"\bmodule loop\b", r"\bmodule\b", r"\bbộ chuyển mạch\b", r"\bbộ nguồn\b"]),
            ("Dây báo cháy / Dây nhiệt", [r"\bdây báo cháy\b", r"\bdây nhiệt\b", r"\bprotectowire\b", r"\blinear heat\b"]),
            ("Đầu báo beam quang học", [r"\bbeam\b", r"\bquang học\b", r"\btia chiếu\b", r"\boptical beam\b"]),
            ("Tủ trung tâm báo cháy", [r"\btrung tâm báo cháy\b", r"\btủ trung tâm\b", r"\bfmz 5000\b", r"\bfmz\b", r"\bfire alarm panel\b", r"\btủ điều khiển chữa cháy\b"]),
            ("Van xả tràn Deluge / Trạm van", [r"\bvan deluge\b", r"\bvan xả tràn\b", r"\bvan tràn\b", r"\bdeluge valve\b", r"\bcuộn hút solenoid\b", r"\bvan wet\b"]),
            ("Bình ắc quy nguồn dự phòng", [r"\bắc quy\b", r"\baccu\b", r"\bbình điện\b", r"\bbattery\b"]),
            ("Cảm biến / Đầu đo khí độc", [r"\bđo khí\b", r"\bnh3\b", r"\bclo\b", r"\bh2\b", r"\bgas detector\b", r"\bkhí độc\b", r"\brò rỉ khí\b"])
        ]

        # 3. Technical Specifications (Quy cách / Công nghệ S)
        self.spec_patterns = [
            ("Địa chỉ (Addressable)", [r"\bđịa chỉ\b", r"\baddressable\b"]),
            ("Thường (Conventional)", [r"\bthường\b", r"\bconventional\b", r"\bvùng\b", r"\bzone\b"]),
            ("Tia chiếu khoảng cách xa (Beam)", [r"\btia chiếu\b", r"\bbeam\b", r"\bkhoảng cách xa\b"]),
            ("Độ nhạy cao / Quang điện tử", [r"\bquang điện\b", r"\bquang học\b", r"\bion\b", r"\bphotoelectric\b"]),
            ("Nhiệt gia tăng / Cố định", [r"\bgia tăng\b", r"\bcố định\b", r"\bđịnh ôn\b"]),
            ("Điện áp 24VDC", [r"\b24vdc\b", r"\b24v\b"]),
            ("Dung lượng cao (>50Ah)", [r"\b100ah\b", r"\b150ah\b", r"\b200ah\b", r"\b12v-100ah\b"]),
            ("Đường kính lớn (DN100 - DN200)", [r"\bdn100\b", r"\bdn150\b", r"\bdn200\b", r"\bdn80\b"]),
            ("Tự động / Xả khí", [r"\btự động\b", r"\bxả khí\b", r"\binterlock\b"])
        ]

        # 4. Context / Environment (Bối cảnh C)
        self.context_patterns = [
            ("Làm việc trên cao (>4m)", [r"\btrên cao\b", r"\bcao >\s*4m\b", r"\b> 4m\b", r"\bgiàn giáo\b"]),
            ("Khu vực gian máy phát điện / Tuabin", [r"\bgian máy\b", r"\btuabin\b", r"\bturbine\b", r"\bmáy phát\b"]),
            ("Trong hầm cáp / Không gian hạn chế", [r"\bhầm cáp\b", r"\bkhông gian hạn chế\b", r"\brãnh cáp\b"]),
            ("Trạm ngoài trời / Hóa chất", [r"\bngoài trời\b", r"\bphòng hóa chất\b", r"\btrạm 110kv\b", r"\btrạm 500kv\b"]),
            ("Phòng điều khiển trung tâm", [r"\bphòng điều khiển\b", r"\btrung tâm\b", r"\bphòng ccr\b"]),
            ("Tiêu chuẩn bình thường", [])
        ]

        # Abstract / Ambiguous Units
        self.ambiguous_units = {"hệ", "hệ thống", "lô", "gói", "trạm", "khối", "khoản", "vụ", "lần"}

    def decompose_vosc(self, text: str) -> Dict[str, str]:
        """Bóc tách chuỗi văn bản thành 4 trường V, O, S, C độc lập."""
        text_lower = " " + text.lower() + " "
        
        # 1. Detect V (Action)
        found_v = []
        for v_name, pats in self.verb_patterns:
            for pat in pats:
                if re.search(pat, text_lower):
                    found_v.append(v_name)
                    break
        v_res = " / ".join(found_v) if found_v else "Chưa xác định hành động"

        # 2. Detect O (Object)
        found_o = []
        for o_name, pats in self.object_patterns:
            for pat in pats:
                if re.search(pat, text_lower):
                    found_o.append(o_name)
                    break
        o_res = " / ".join(found_o) if found_o else "Thiết bị chưa phân loại / Khái quát"

        # 3. Detect S (Spec)
        found_s = []
        for s_name, pats in self.spec_patterns:
            for pat in pats:
                if re.search(pat, text_lower):
                    found_s.append(s_name)
                    break
        s_res = " / ".join(found_s) if found_s else "Tiêu chuẩn / Không nêu quy cách"

        # 4. Detect C (Context)
        found_c = []
        for c_name, pats in self.context_patterns:
            for pat in pats:
                if re.search(pat, text_lower):
                    found_c.append(c_name)
                    break
        c_res = " / ".join(found_c) if found_c else "Điều kiện mặt bằng thông thường"

        return {
            "v_action": v_res,
            "o_object": o_res,
            "s_spec": s_res,
            "c_context": c_res
        }

    def infer_technical_nature(self, vosc: Dict[str, str], raw_unit: str) -> Dict[str, str]:
        """
        Suy luận kỹ thuật:
        - Làm cái gì? (What)
        - Làm như thế nào? (How)
        """
        v = vosc["v_action"]
        o = vosc["o_object"]
        s = vosc["s_spec"]
        
        # 1. Suy luận "Làm cái gì?"
        if "Lắp mới" in v:
            what = f"Lắp đặt bổ sung mới hoặc thay thế toàn diện thiết bị [{o}] ({s}), bao gồm cung cấp vật tư chính và phụ kiện."
        elif "Thay thế linh kiện" in v:
            what = f"Tháo dỡ linh kiện cũ hỏng và thay thế linh kiện mới cho [{o}] ({s})."
        elif "Bảo dưỡng" in v or "Đo kiểm" in v or "Hiệu chỉnh" in v:
            what = f"Bảo dưỡng định kỳ, vệ sinh cơ khí, đo kiểm thông mạch và hiệu chỉnh tín hiệu kỹ thuật cho [{o}] ({s}), KHÔNG thay mới thiết bị chính."
        elif "Tháo dỡ" in v:
            what = f"Tháo dỡ, thu hồi thiết bị [{o}] khỏi vị trí lắp đặt hiện hữu."
        else:
            what = f"Thực hiện công việc liên quan đến [{o}] nhưng chưa rõ bản chất bảo dưỡng hay xây lắp mới."

        # 2. Suy luận "Làm như thế nào?"
        steps_list = []
        if "Bảo dưỡng" in v or "Đo kiểm" in v or "Hiệu chỉnh" in v:
            steps_list.append("1. Khảo sát hiện trường, cô lập hệ thống tín hiệu an toàn.")
            steps_list.append("2. Tháo nắp, thổi bụi, xịt dung dịch bảo dưỡng chuyên dụng (RP7/Cồn).")
            steps_list.append("3. Dùng đồng hồ VOM/thiết bị tạo khói/nhiệt chuyên dụng đo kiểm tra thông số tín hiệu.")
            steps_list.append("4. Hiệu chỉnh ngưỡng đáp ứng và thử nghiệm kích hoạt liên động.")
            steps_list.append("5. Khôi phục trạng thái hoạt động bình thường của hệ thống.")
            how = " ".join(steps_list)
        elif "Lắp mới" in v or "Thay thế" in v:
            steps_list.append("1. Đo đạc lấy dấu, định vị giá đỡ/vị trí lắp đặt thiết bị.")
            steps_list.append("2. Lắp cố định đế và thiết bị bằng đinh vít nở.")
            steps_list.append("3. Đấu nối cáp tín hiệu/cáp nguồn vào cầu đấu thiết bị.")
            steps_list.append("4. Cài đặt địa chỉ số/cấu hình thông số trên phần mềm tủ trung tâm.")
            steps_list.append("5. Thử nghiệm nghiệm thu bàn giao đưa vào vận hành.")
            how = " ".join(steps_list)
        else:
            how = "Chưa đủ dữ kiện quy trình để suy luận các bước thi công."

        return {"what": what, "how": how}

    def evaluate_audit_warning(self, vosc: Dict[str, str], raw_unit: str, raw_vol: float, raw_name: str) -> Dict[str, str]:
        """
        Đánh giá tính minh bạch hồ sơ và cảnh báo:
        - level: GREEN (Rõ ràng), YELLOW (Cần làm rõ biến thể/quy đổi), RED (Mập mờ/Không đủ điều kiện)
        - can_lam_ro: Nội dung cụ thể cần làm rõ
        """
        v = vosc["v_action"]
        o = vosc["o_object"]
        unit_lower = str(raw_unit).strip().lower()

        # 1. KIỂM TRA MỨC ĐỘ ĐỎ (RED) - MẬP MỜ / CẤM ÉP ĐỊNH MỨC
        if unit_lower in self.ambiguous_units:
            return {
                "level": "RED",
                "status_text": "🔴 MẬP MỜ (ĐVT PHI ĐỊNH MỨC)",
                "can_lam_ro": f"Cần làm rõ: ĐVT '{raw_unit}' là đơn vị trừu tượng (Hệ/Gói/Lô). Phải yêu cầu bóc tách chi tiết danh mục gồm cụ thể bao nhiêu thiết bị, model gì, số lượng từng cái; TUYỆT ĐỐI KHÔNG ÉP ĐỊNH MỨC ĐƠN CHIẾC."
            }

        if o == "Thiết bị chưa phân loại / Khái quát":
            return {
                "level": "RED",
                "status_text": "🔴 MẬP MỜ (THIẾU ĐỐI TƯỢNG)",
                "can_lam_ro": "Cần làm rõ: Tên công việc ghi khái quát, không xác định được thiết bị kỹ thuật cụ thể (là đầu báo, chuông đèn, van hay tủ?). Không đủ dữ kiện áp định mức."
            }

        # 2. KIỂM TRA MỨC ĐỘ VÀNG (YELLOW) - MÂU THUẪN HOẶC CẦN BIẾN THỂ / QUY ĐỔI
        clarifications = []
        
        # Kiểm tra xung đột: Tên ghi bảo dưỡng nhưng đối chiếu sang mã lắp mới
        if "Bảo dưỡng" in v or "Đo kiểm" in v or "Hiệu chỉnh" in v:
            clarifications.append("Có mua mới thay thế linh kiện/thiết bị không? (Nếu chỉ bảo dưỡng đo kiểm thì bắt buộc áp Biến thể A1 hoặc mã E6C để loại trừ chi phí vật tư đế/thiết bị).")

        # Kiểm tra quy đổi ĐVT
        if any(term in o for term in ["Đầu báo khói", "Đầu báo nhiệt", "Đầu báo"]) and unit_lower in ["bộ", "cái", "đầu"]:
            clarifications.append("ĐVT thiết kế là 'Cái/Bộ' nhưng ĐM gốc quy định '10 đầu' (Cần áp dụng hệ số quy đổi 0,1 để tránh nhân đội khối lượng 10 lần).")

        if clarifications:
            return {
                "level": "YELLOW",
                "status_text": "🟡 CẦN LÀM RÕ ĐIỀU KIỆN ÁP DỤNG",
                "can_lam_ro": "Cần làm rõ: " + " | ".join(clarifications)
            }

        # 3. MỨC ĐỘ XANH (GREEN) - ĐẠT CHUẨN
        return {
            "level": "GREEN",
            "status_text": "🟢 ĐỦ ĐIỀU KIỆN ÁP ĐỊNH MỨC",
            "can_lam_ro": "Đủ điều kiện áp định mức (Thông tin rõ ràng về hành động, thiết bị và đơn vị tính)."
        }

    def match_proposed_norm_details(self, vosc: Dict[str, str], raw_unit: str, raw_vol: float) -> Dict[str, Any]:
        """
        Khớp mã định mức và trích xuất BẠCH HÓA NỘI HÀM ĐỊNH MỨC GỐC từ CSDL:
        - Tên ĐM gốc & Pháp lý
        - ĐVT gốc
        - Thành phần công việc gốc (các bước)
        - Hao phí nhân công (tổng công & bậc thợ)
        - Hao phí máy
        - ĐVT & Khối lượng quy đổi
        """
        v = vosc["v_action"]
        o = vosc["o_object"]
        
        # Logic ánh xạ mã định mức
        matched_code = ""
        is_maintenance = "Bảo dưỡng" in v or "Đo kiểm" in v or "Hiệu chỉnh" in v

        if "Đầu báo khói" in o or "Đầu báo nhiệt" in o or "Đầu báo" in o:
            matched_code = "E6C.3.02" if is_maintenance else "BD.41111"
        elif "Đèn báo cháy" in o:
            matched_code = "BD.41121"
        elif "Nút nhấn khẩn" in o:
            matched_code = "E6C.3.01" if is_maintenance else "BD.41131"
        elif "Chuông" in o:
            matched_code = "BD.41141"
        elif "Module" in o or "Bộ chuyển mạch" in o:
            matched_code = "E6C.3.03" if is_maintenance else "BD.42201"
        elif "Dây báo cháy" in o or "Dây nhiệt" in o:
            matched_code = "E6C.3.03"
        elif "Đầu báo beam" in o:
            matched_code = "E6C.3.04"
        elif "Tủ trung tâm" in o:
            matched_code = "E6C.3.05"
        elif "Van xả tràn" in o or "Van deluge" in o:
            matched_code = "E6C.3.06"
        elif "Bình ắc quy" in o:
            matched_code = "E6B.1.13.01"
        elif "Cảm biến" in o or "khí độc" in o:
            matched_code = "E6C.1.31"
        else:
            matched_code = "BD.41111" # Mặc định kiểm tra

        norm_data = self.norm_manager.get_norm(matched_code)
        if not norm_data:
            return {
                "norm_code": matched_code,
                "norm_name": "Chưa có trong CSDL định mức EVN",
                "norm_unit": "-",
                "legal_basis": "-",
                "steps_summary": "-",
                "labor_summary": "-",
                "machine_summary": "-",
                "conv_factor": 1.0,
                "converted_vol": raw_vol
            }

        # 1. Tên & Pháp lý
        n_name = norm_data.get("name", "")
        n_basis = norm_data.get("legal_basis", "")
        n_unit = norm_data.get("unit", "")

        # 2. Thành phần công việc gốc (Quy trình các bước)
        w_steps = norm_data.get("work_steps", [])
        if w_steps:
            steps_desc = "; ".join([f"{st.get('name', '')} ({st.get('pct', '')})" for st in w_steps])
        else:
            steps_desc = "Theo quy trình bảo dưỡng tiêu chuẩn của EVN."

        # 3. Hao phí nhân công chi tiết
        r_labor = norm_data.get("resources", {}).get("labor", [])
        tot_labor_rate = sum([lb.get("rate", 0.0) for lb in r_labor])
        labor_details = " + ".join([f"{lb.get('rate', 0.0)} công {lb.get('name', '')} ({lb.get('spec', '')})" for lb in r_labor])
        labor_summary = f"Tổng: {round(tot_labor_rate, 4)} công/{n_unit} [{labor_details}]" if labor_details else f"{tot_labor_rate} công/{n_unit}"

        # 4. Hao phí máy thi công
        r_mach = norm_data.get("resources", {}).get("machines", [])
        if r_mach:
            machine_summary = " + ".join([f"{m.get('rate', 0.0)} ca {m.get('name', '')}" for m in r_mach])
        else:
            machine_summary = "Dụng cụ cầm tay thông dụng (Không tính ca máy lớn)"

        # 5. Hệ số quy đổi
        conv_factor = 1.0
        u_raw_clean = str(raw_unit).lower().strip()
        if "10" in n_unit and "10" not in u_raw_clean:
            conv_factor = 0.1
        elif "5" in n_unit and "5" not in u_raw_clean:
            conv_factor = 0.2

        converted_vol = round(float(raw_vol or 0.0) * conv_factor, 4)

        return {
            "norm_code": matched_code,
            "norm_name": n_name,
            "norm_unit": n_unit,
            "legal_basis": n_basis,
            "steps_summary": steps_desc,
            "labor_summary": labor_summary,
            "machine_summary": machine_summary,
            "conv_factor": conv_factor,
            "converted_vol": converted_vol
        }

    def audit_raw_dataset(self, raw_items: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Xử lý toàn diện danh mục thô:
        - Bóc tách V-O-S-C (4 cột riêng)
        - Suy luận kỹ thuật
        - Đánh giá cảnh báo 'Cần làm rõ gì?'
        - Bạch hóa định mức đề xuất
        - Gom cụm mẫu lặp & đếm tần suất
        """
        audited_rows = []
        cluster_signatures = []

        # Bước 1: Phân tích từng dòng độc lập
        for item in raw_items:
            stt = item.get("stt", len(audited_rows) + 1)
            raw_name = str(item.get("job_name", "")).strip()
            raw_desc = str(item.get("description", raw_name)).strip()
            raw_unit = str(item.get("unit", "Bộ")).strip()
            try:
                raw_vol = float(item.get("volume", 1.0) or 1.0)
            except:
                raw_vol = 1.0

            combined_text = f"{raw_name} {raw_desc}"
            
            # 1. Phân rã V-O-S-C
            vosc = self.decompose_vosc(combined_text)

            # 2. Suy luận bản chất
            nature = self.infer_technical_nature(vosc, raw_unit)

            # 3. Cảnh báo "Cần làm rõ gì?"
            audit_flag = self.evaluate_audit_warning(vosc, raw_unit, raw_vol, raw_name)

            # 4. Bạch hóa định mức đề xuất
            norm_details = self.match_proposed_norm_details(vosc, raw_unit, raw_vol)

            # 5. Xác định cụm mẫu lặp (Pattern Signature)
            # Cụm được định danh bằng cặp [Hành động V + Đối tượng O]
            cluster_id = f"CỤM: [{vosc['v_action']}] - [{vosc['o_object']}]"
            cluster_signatures.append(cluster_id)

            row_data = {
                "stt": stt,
                "raw_job_name": raw_name,
                "raw_desc": raw_desc,
                "raw_unit": raw_unit,
                "raw_volume": raw_vol,
                
                # 4 Cột độc lập để filter
                "v_action": vosc["v_action"],
                "o_object": vosc["o_object"],
                "s_spec": vosc["s_spec"],
                "c_context": vosc["c_context"],

                # Suy luận bản chất
                "nature_what": nature["what"],
                "nature_how": nature["how"],

                # Cảnh báo trọng tâm
                "audit_level": audit_flag["level"],
                "audit_status": audit_flag["status_text"],
                "can_lam_ro": audit_flag["can_lam_ro"],

                # Nội hàm định mức đề xuất
                "prop_norm_code": norm_details["norm_code"],
                "prop_norm_name": norm_details["norm_name"],
                "prop_norm_unit": norm_details["norm_unit"],
                "prop_legal_basis": norm_details["legal_basis"],
                "prop_work_steps": norm_details["steps_summary"],
                "prop_labor_rate": norm_details["labor_summary"],
                "prop_machine_rate": norm_details["machine_summary"],

                # Quy đổi
                "conv_factor": norm_details["conv_factor"],
                "converted_vol": norm_details["converted_vol"],

                # Cụm mẫu lặp (sẽ cập nhật count ở bước 2)
                "cluster_id": cluster_id,
                "cluster_count": 1
            }
            audited_rows.append(row_data)

        # Bước 2: Thống kê tần suất gom cụm mẫu lặp
        cluster_counter = Counter(cluster_signatures)
        for r in audited_rows:
            r["cluster_count"] = cluster_counter[r["cluster_id"]]

        # Bước 3: Thống kê tổng hợp hồ sơ
        level_counter = Counter([r["audit_level"] for r in audited_rows])
        total_rows = len(audited_rows)
        summary_stats = {
            "total_items": total_rows,
            "green_count": level_counter.get("GREEN", 0),
            "yellow_count": level_counter.get("YELLOW", 0),
            "red_count": level_counter.get("RED", 0),
            "clusters": [
                {"cluster_name": c_name, "count": cnt, "pct": round(cnt / total_rows * 100, 1) if total_rows else 0}
                for c_name, cnt in cluster_counter.most_common()
            ]
        }

        return {
            "audited_rows": audited_rows,
            "summary_stats": summary_stats
        }
