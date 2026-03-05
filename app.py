import streamlit as st
from pdf2docx import Converter
from docx import Document
import pytesseract
from PIL import Image
import os, tempfile
from pdf2image import convert_from_path

# Cấu hình giao diện Neon thu gọn
st.set_page_config(page_title="Học Liệu Số Pro", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #000b14; color: #00d4ff; }
    .main-box {
        border: 1px solid #00d4ff; border-radius: 10px;
        padding: 10px; background-color: rgba(0, 20, 40, 0.8);
        box-shadow: 0 0 8px #00d4ff;
    }
    .stButton>button { width: 100%; border: 1px solid #00d4ff; background: transparent; color: #00d4ff; }
    .stButton>button:hover { background: #00d4ff; color: #000; box-shadow: 0 0 15px #00d4ff; }
    </style>
    """, unsafe_allow_html=True)

# Hàm hỗ trợ phân tách khoảng trang (Ví dụ: "1-50" -> start=0, end=50)
def parse_pages(range_str):
    try:
        if "-" in range_str:
            s, e = map(int, range_str.split("-"))
            return s - 1, e
        return int(range_str) - 1, int(range_str)
    except:
        return 0, None

st.markdown('<div style="border: 1px solid #00d4ff; padding: 5px; text-align: center; border-radius: 5px; margin-bottom: 10px;">'
            '<h2 style="margin:0; color:#00d4ff;">🚀 CHUYỂN ĐỔI THEO PHÂN ĐOẠN</h2></div>', unsafe_allow_html=True)

c1, c2 = st.columns([1, 1])

with c1:
    st.markdown('<div class="main-box">', unsafe_allow_html=True)
    st.write("📂 **CÀI ĐẶT**")
    f = st.file_uploader("Tải lên PDF", type="pdf", label_visibility="collapsed")
    is_ocr = st.toggle("Chế độ quét ảnh (OCR Scan)", value=False)
    pg_input = st.text_input("Nhập trang cần lấy (Ví dụ: 1-50 hoặc 5)", help="Để trống để chuyển toàn bộ")
    st.markdown('</div>', unsafe_allow_html=True)

with c2:
    st.markdown('<div class="main-box">', unsafe_allow_html=True)
    st.write("⚡ **TRẠNG THÁI**")
    
    if f:
        if st.button("BẮT ĐẦU TRÍCH XUẤT"):
            start, end = parse_pages(pg_input)
            out_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
            
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp.write(f.getvalue())
                    tmp_p = tmp.name

                if not is_ocr:
                    # Chế độ thông thường: Dùng start/end của pdf2docx
                    with st.spinner(f"Đang chuyển trang {start+1} đến {end if end else 'hết'}..."):
                        cv = Converter(tmp_p)
                        cv.convert(out_file, start=start, end=end)
                        cv.close()
                else:
                    # Chế độ OCR: Dùng first_page/last_page của pdf2image
                    with st.spinner("Đang quét OCR từng trang..."):
                        # Giảm DPI xuống 150 để tiết kiệm RAM cho file lớn
                        imgs = convert_from_path(tmp_p, dpi=150, first_page=start+1, last_page=end)
                        doc = Document()
                        progress_bar = st.progress(0)
                        for i, img in enumerate(imgs):
                            txt = pytesseract.image_to_string(img, lang='vie')
                            doc.add_paragraph(f"--- TRANG {start + i + 1} ---")
                            doc.add_paragraph(txt)
                            progress_bar.progress((i + 1) / len(imgs))
                        doc.save(out_file)
                
                st.success("Đã hoàn thành phân đoạn!")
                with open(out_file, "rb") as dl:
                    st.download_button("📥 TẢI ĐOẠN VĂN BẢN NÀY", dl, file_name=f"Trang_{pg_input}.docx")
                
                os.remove(tmp_p)
                os.remove(out_file)
            except Exception as e:
                st.error("Lỗi: Kiểm tra lại số trang hoặc dung lượng file.")
    else:
        st.info("Chưa có file nào được tải lên.")
    st.markdown('</div>', unsafe_allow_html=True)
