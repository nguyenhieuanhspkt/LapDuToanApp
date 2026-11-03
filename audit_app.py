"""
Streamlit Standalone Web App:
Module Tiền Thẩm Tra & Bóc Tách Bản Chất Kỹ Thuật Dữ Liệu Thô PAKT
Hoàn toàn độc lập với module tính toán dự toán.
Author: Tổ Thẩm Định EVN
"""

import streamlit as st
import pandas as pd
import re
from pathlib import Path
import sys

# Ensure project root is in sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from core.semantic_auditor import SemanticAuditor
from core.audit_excel_exporter import AuditExcelExporter

st.set_page_config(
    page_title="Thẩm Tra & Bóc Tách Dữ Liệu Thô PAKT - EVN",
    page_icon="🔬",
    layout="wide"
)

# Custom Styling EVN Navy
st.markdown("""
<style>
    .main-title {
        font-size: 26px;
        font-weight: 700;
        color: #002060;
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 14px;
        color: #555;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #F8F9FA;
        border-top: 4px solid #002060;
        padding: 12px;
        border-radius: 6px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🔬 HỆ THỐNG THẨM TRA & BÓC TÁCH DỮ LIỆU THÔ PAKT (ĐỘC LẬP)</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Phân tích ngữ nghĩa 4 thành phần V-O-S-C, cảnh báo mập mờ, gom cụm mẫu lặp và bạch hóa nội hàm định mức đề xuất trước khi đưa vào lập dự toán</div>', unsafe_allow_html=True)

# Khởi tạo auditor & exporter
if "auditor" not in st.session_state:
    st.session_state.auditor = SemanticAuditor()
if "exporter" not in st.session_state:
    st.session_state.exporter = AuditExcelExporter()

# Sidebar: Cấu hình thông tin dự án
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1d/EVN_Logo.svg/300px-EVN_Logo.svg.png", width=110)
    st.markdown("### 🏢 HỒ SƠ THẨM TRA")
    proj_title = st.text_input("Tên dự án / Công trình:", value="BẢO DƯỠNG PCCC NĂM 2026 - NMNĐ VĨNH TÂN 4")
    st.divider()
    st.markdown("### ℹ️ NGUYÊN TẮC THẨM TRA")
    st.info("""
    - **Cột riêng V-O-S-C**: Cho phép lọc theo Hành động, Đối tượng, Quy cách, Bối cảnh.
    - **Cảnh báo trọng tâm**: Chỉ nêu rõ *'Cần làm rõ gì?'*.
    - **Bạch hóa định mức**: Trích xuất tên ĐM gốc, quy trình các bước, hao phí nhân công và máy thi công.
    - **Tuyệt đối không ép mã** nếu đơn vị tính là 'Hệ'/'Gói'/'Lô'.
    """)

# Vùng nhập liệu thô
st.markdown("#### 📥 1. Nhập danh mục dữ liệu thô từ PAKT / Thiết kế")
col_input_mode, col_helper = st.columns([8, 4])
with col_helper:
    st.caption("Định dạng cột ưu tiên khi copy từ Excel:")
    st.code("STT\tTên công việc\tMô tả/Quy cách\tĐVT\tKhối lượng", language="text")

default_demo_text = """1\tThay mới lắp đặt thiết bị đầu báo khói địa chỉ\tĐầu báo khói quang điện tử địa chỉ Hochiki ALN-V gian máy\tBộ\t20
2\tBảo dưỡng, vệ sinh thổi bụi đầu báo khói\tVệ sinh buồng khói, xịt cồn bảo dưỡng định kỳ khu nhà phụ trợ\tCái\t50
3\tKiểm tra, thử nghiệm độ nhạy đầu báo khói địa chỉ\tDùng bình tạo khói chuyên dụng thử kích hoạt tín hiệu loop\tCái\t35
4\tBảo dưỡng hệ thống PCCC gian máy chính\tBảo dưỡng toàn bộ thiết bị PCCC khu vực tuabin máy phát\tHệ\t1
5\tThí nghiệm tác động cuộn hút van Deluge xả tràn\tĐo kiểm cuộn hút solenoid van tràn Deluge DN100 hệ chữa cháy Foam\tCái\t6
6\tKiểm tra đo nội trở bình ắc quy dự phòng tủ PCCC\tĐo nội trở, dung lượng Ah bình ắc quy 24VDC-100Ah trạm 110kV\tBình\t12
7\tBảo dưỡng thiết bị phụ trợ linh tinh\tKhông có thông số cụ thể\tGói\t1"""

raw_input_text = st.text_area(
    "Dán dữ liệu đa cột từ file Excel dự toán / PAKT vào đây:",
    value=default_demo_text,
    height=160
)

def parse_paste_to_items(text: str):
    items = []
    lines = [l.strip() for l in text.strip().split("\n") if l.strip()]
    for idx, line in enumerate(lines, start=1):
        parts = [p.strip() for p in re.split(r"\t+| {2,}", line) if p.strip()]
        if not parts:
            continue
        
        # Nhận diện STT
        stt = idx
        if parts[0].isdigit():
            stt = int(parts[0])
            rem = parts[1:]
        else:
            rem = parts[:]

        # Nhận diện số lượng (phần tử số cuối cùng)
        vol = 1.0
        for p in reversed(list(rem)):
            p_clean = re.sub(r'[^0-9.,]', '', p).replace(',', '.')
            try:
                vol = float(p_clean)
                rem.remove(p)
                break
            except:
                pass

        # Nhận diện ĐVT
        known_units = {'bộ', 'cái', 'đầu', 'bình', 'trung tâm', 'hệ thống', 'tủ', 'gói', 'm', 'mét', 'kg', 'tấm', 'cuộn', 'điểm', 'trạm', 'hệ', 'lô'}
        unit = "Bộ"
        for p in list(rem):
            if p.lower() in known_units:
                unit = p
                rem.remove(p)
                break

        name = rem[0] if rem else f"Công tác {stt}"
        desc = rem[1] if len(rem) > 1 else name

        items.append({
            "stt": stt,
            "job_name": name,
            "description": desc,
            "unit": unit,
            "volume": vol
        })
    return items

if st.button("🚀 BẮT ĐẦU PHÂN TÍCH NGỮ NGHĨA V-O-S-C & THẨM TRA", type="primary", use_container_width=True):
    raw_items = parse_paste_to_items(raw_input_text)
    if not raw_items:
        st.warning("Không tìm thấy dữ liệu hợp lệ để phân tích.")
    else:
        st.session_state["audit_result"] = st.session_state.auditor.audit_raw_dataset(raw_items)
        st.session_state["audit_project_title"] = proj_title

# Hiển thị kết quả phân tích
if "audit_result" in st.session_state:
    res = st.session_state["audit_result"]
    rows = res["audited_rows"]
    stats = res["summary_stats"]

    st.divider()
    st.markdown("### 📊 2. Bảng Điều Khiển Chất Lượng Hồ Sơ Thô")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("Tổng số dòng thô", f"{stats['total_items']} dòng")
    with c2:
        st.metric("🟢 Đạt chuẩn (Rõ ràng)", f"{stats['green_count']} dòng")
    with c3:
        st.metric("🟡 Cần làm rõ điều kiện", f"{stats['yellow_count']} dòng")
    with c4:
        st.metric("🔴 Mập mờ (Cấm ép mã)", f"{stats['red_count']} dòng")
    with c5:
        st.metric("Cụm mẫu lặp phát hiện", f"{len(stats['clusters'])} cụm")

    # Nút tải file Excel
    st.markdown("#### 📥 3. Trích xuất Báo Cáo Thẩm Tra Excel")
    excel_bytes = st.session_state.exporter.export_workbook(res, project_title=st.session_state.get("audit_project_title", "PAKT"))
    st.download_button(
        label="📄 TẢI FILE EXCEL BÁO CÁO THẨM TRA DỮ LIỆU THÔ (AUTOFILTER + VOSC + NỘI HÀM ĐỊNH MỨC)",
        data=excel_bytes,
        file_name="BaoCao_ThamTra_DuLieuTho_PAKT_VOSC.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

    # Hiển thị Tabs chi tiết
    t_vosc, t_clusters = st.tabs(["📋 BẢNG THẨM TRA CHI TIẾT (V-O-S-C & CẢNH BÁO)", "🔄 THỐNG KÊ CỤM MẪU LẶP LẠI"])

    with t_vosc:
        st.caption("💡 Các cột V, O, S, C được tách độc lập để lọc. Cột 'CẦN LÀM RÕ GÌ?' chỉ rõ hành động cần yêu cầu bên lập hồ sơ.")
        
        display_data = []
        for r in rows:
            display_data.append({
                "STT": r["stt"],
                "Tên thô PAKT": r["raw_job_name"],
                "ĐVT": r["raw_unit"],
                "KL": r["raw_volume"],
                "V (Hành động)": r["v_action"],
                "O (Đối tượng)": r["o_object"],
                "S (Quy cách)": r["s_spec"],
                "C (Bối cảnh)": r["c_context"],
                "CẦN LÀM RÕ GÌ?": r["can_lam_ro"],
                "Trạng thái": r["audit_status"],
                "Mã ĐM": r["prop_norm_code"],
                "Tên ĐM gốc": r["prop_norm_name"],
                "ĐVT ĐM": r["prop_norm_unit"],
                "Quy trình các bước": r["prop_work_steps"],
                "Hao phí nhân công gốc": r["prop_labor_rate"],
                "Hệ số quy đổi": r["conv_factor"],
                "KL quy đổi": r["converted_vol"],
                "Cụm mẫu lặp": r["cluster_id"],
                "Số dòng cùng cụm": r["cluster_count"]
            })

        df_display = pd.DataFrame(display_data)

        # Bộ lọc tương tác trực tiếp trên giao diện
        col_fil1, col_fil2, col_fil3 = st.columns(3)
        with col_fil1:
            all_v = ["Tất cả"] + list(df_display["V (Hành động)"].unique())
            sel_v = st.selectbox("Lọc theo Hành động (V):", options=all_v)
        with col_fil2:
            all_o = ["Tất cả"] + list(df_display["O (Đối tượng)"].unique())
            sel_o = st.selectbox("Lọc theo Đối tượng (O):", options=all_o)
        with col_fil3:
            all_st = ["Tất cả"] + list(df_display["Trạng thái"].unique())
            sel_st = st.selectbox("Lọc theo Trạng thái:", options=all_st)

        filtered_df = df_display.copy()
        if sel_v != "Tất cả":
            filtered_df = filtered_df[filtered_df["V (Hành động)"] == sel_v]
        if sel_o != "Tất cả":
            filtered_df = filtered_df[filtered_df["O (Đối tượng)"] == sel_o]
        if sel_st != "Tất cả":
            filtered_df = filtered_df[filtered_df["Trạng thái"] == sel_st]

        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True
        )

    with t_clusters:
        st.caption("Danh sách các nhóm công việc có cùng cấu trúc ngữ nghĩa (Hành động + Đối tượng) và tần suất xuất hiện:")
        df_clusters = pd.DataFrame(stats["clusters"])
        df_clusters.columns = ["Cụm Mẫu Lặp", "Số Dòng Xuất Hiện", "Tỷ Lệ (%)"]
        st.dataframe(df_clusters, use_container_width=True, hide_index=True)
