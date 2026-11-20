import os
import sys
import tempfile
import io
from pathlib import Path
import streamlit as st
import pandas as pd

# Add src to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from c12_tool.parser import parse_c12_records, C12ParseError
from c12_tool.excel_writer import export_records

# Cấu hình trang
st.set_page_config(
    page_title="Xuất Dữ Liệu Thông Báo BHXH (C12)",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS cho giao diện hiện đại, chuyên nghiệp
st.markdown(
    """
    <style>
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 5px;
    }
    .sub-header {
        font-size: 14px;
        color: #64748B;
        margin-bottom: 25px;
    }
    .stDownloadButton button {
        background-color: #10B981 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        padding: 10px 24px !important;
        font-size: 16px !important;
        width: 100%;
        border: none !important;
    }
    .stDownloadButton button:hover {
        background-color: #059669 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Tiêu đề ứng dụng
st.markdown('<div class="main-header">📑 HỆ THỐNG XỬ LÝ & XUẤT THÔNG BÁO BHXH C12</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">Tự động đọc dữ liệu C12 (Việt Nam & Nước Ngoài), trích xuất chỉ tiêu B, Quỹ lương, C (Đã nộp) và điền chuẩn vào mẫu Excel.</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("⚙️ Hướng dẫn & Tùy chọn")
    st.info(
        "1. Kéo thả hoặc chọn các file **PDF / Hình ảnh** thông báo C12 vào khung bên phải.\n"
        "2. Bấm **Bắt đầu xử lý** để hệ thống tự động nhận diện và trích xuất số liệu.\n"
        "3. Xem trước bảng kết quả và bấm **Tải file Excel**."
    )
    st.markdown("---")
    st.markdown("**Đặc tính nổi bật:**")
    st.markdown("- ✅ Tự động phân loại đơn vị Việt Nam (**YN**) và Nước ngoài (**IC**)")
    st.markdown("- ✅ Tự động khớp tháng 1 - 12")
    st.markdown("- ✅ Điền đúng số tiền mang sang và số tiền đã nộp")
    st.markdown("- ✅ Chạy mượt mà trực tiếp trên nền web")

# Khu vực Upload tập tin
uploaded_files = st.file_uploader(
    "📤 Chọn hoặc kéo thả các tập tin C12 (PDF hoặc Ảnh scan):",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
    help="Có thể chọn nhiều file cùng lúc của nhiều tháng khác nhau.",
)

if uploaded_files:
    st.success(f"Đã chọn **{len(uploaded_files)}** tập tin sẵn sàng xử lý.")
    
    col_btn, col_empty = st.columns([1, 3])
    with col_btn:
        start_process = st.button("🚀 Bắt đầu xử lý dữ liệu", type="primary", use_container_width=True)

    if start_process:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        all_records = []
        errors = []
        
        # Tạo thư mục tạm để lưu file upload cho parser xử lý
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            total = len(uploaded_files)
            
            for idx, uploaded_file in enumerate(uploaded_files):
                status_text.text(f"Đang xử lý [{idx+1}/{total}]: {uploaded_file.name}...")
                file_dest = tmp_path / uploaded_file.name
                with open(file_dest, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                try:
                    recs = parse_c12_records(file_dest)
                    all_records.extend(recs)
                except Exception as e:
                    errors.append(f"{uploaded_file.name}: {e}")
                
                progress_bar.progress(int((idx + 1) / total * 100))
        
        status_text.empty()
        progress_bar.empty()
        
        if errors:
            with st.expander("⚠️ Cảnh báo một số file không đọc được", expanded=False):
                for err in errors:
                    st.error(err)
        
        if not all_records:
            st.error("Không trích xuất được bản ghi C12 nào từ các tập tin đã tải lên.")
        else:
            st.success(f"🎉 Xử lý thành công **{len(all_records)}** thông báo C12!")
            
            # Xuất ra file Excel trong bộ nhớ
            template_path = Path(__file__).resolve().parent / "THÔNG BÁO BXH.xlsx"
            
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_out:
                out_path = Path(tmp_out.name)
            
            try:
                export_records(all_records, out_path, template_path=template_path if template_path.is_file() else None)
                with open(out_path, "rb") as f:
                    excel_bytes = f.read()
            finally:
                if out_path.exists():
                    try:
                        os.unlink(out_path)
                    except Exception:
                        pass
            
            # Hiển thị bảng tổng hợp kết quả xem trước
            st.markdown("### 📊 Xem trước số liệu đã trích xuất")
            
            yn_data = [
                {
                    "Tháng": r.month,
                    "Phát sinh trong kỳ (B)": f"{r.b_total:,.0f}" if r.b_total is not None else "-",
                    "Quỹ lương (2)": f"{r.row2_bhtnld:,.0f}" if r.row2_bhtnld is not None else "-",
                    "Số đã nộp trong kỳ (C)": f"{r.c_total:,.0f}" if r.c_total is not None else "-",
                    "File nguồn": r.pdf_path.name,
                }
                for r in all_records if r.unit == "YN"
            ]
            
            ic_data = [
                {
                    "Tháng": r.month,
                    "Phát sinh trong kỳ (B)": f"{r.b_total:,.0f}" if r.b_total is not None else "-",
                    "Quỹ lương (2)": f"{r.row2_bhtnld:,.0f}" if r.row2_bhtnld is not None else "-",
                    "Số đã nộp trong kỳ (C)": f"{r.c_total:,.0f}" if r.c_total is not None else "-",
                    "File nguồn": r.pdf_path.name,
                }
                for r in all_records if r.unit == "IC"
            ]
            
            tab_yn, tab_ic = st.tabs(["🇻🇳 Đơn vị YN (Việt Nam)", "🌍 Đơn vị IC (Nước Ngoài)"])
            with tab_yn:
                if yn_data:
                    df_yn = pd.DataFrame(yn_data).sort_values("Tháng")
                    st.dataframe(df_yn, use_container_width=True, hide_index=True)
                else:
                    st.info("Chưa có dữ liệu đơn vị YN.")
            
            with tab_ic:
                if ic_data:
                    df_ic = pd.DataFrame(ic_data).sort_values("Tháng")
                    st.dataframe(df_ic, use_container_width=True, hide_index=True)
                else:
                    st.info("Chưa có dữ liệu đơn vị IC.")
            
            st.markdown("---")
            st.download_button(
                label="📥 TẢI FILE EXCEL KẾT QUẢ (.XLSX)",
                data=excel_bytes,
                file_name="THONG_BAO_BHXH_TONG_HOP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
else:
    st.info("👆 Vui lòng tải các file C12 lên để bắt đầu.")
