import streamlit as st
from pdf2docx import Converter
from docx import Document
import pytesseract
from PIL import Image
import os, tempfile, gc
from pdf2image import convert_from_path

st.set_page_config(page_title="Học Liệu Số Pro", layout="wide")

# CSS tối giản giao diện trong 1 màn hình
st.markdown("""
    <style>
    .stApp { background-color: #000b14; color: #00d4ff; font-size: 14px; }
    .main-box {
        border: 1px solid #00d4ff; border-radius: 8px;
        padding: 15px; background-color: rgba(0, 20, 40, 0.9);
        box-shadow: 0 0 10px #00d4ff;
    }
    .stButton>button { width: 100%; border: 1px solid #00d4ff; background: transparent; color: #00d4ff; font-weight: bold; }
    .stButton>button:hover { background: #00d4ff; color: #000; box-shadow: 0 0 20px #00d4ff; }
    input, .stTextInput>div>div>input { background-color: #001a33 !important; color: #00d4ff !important; border: 1px solid #00d4ff !important; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<h2 style="text-align:center; color:#00d4ff; border: 1px solid #00d4ff; padding:10px; border-radius:10px;">🚀 HỆ THỐNG CHUYỂN ĐỔI HỌC LIỆU SỐ</h2>', unsafe_allow_html=True)

# Chia cột để tất cả nằm trong 1 màn hình
col1, col2 = st.columns([1, 1.2])

with col1:
    st.markdown('<div class="main-box">', unsafe_allow_html=True)
    st.markdown("### 📥 ĐẦU VÀO")
    mode = st.radio("Chế độ xử lý:", ["PDF sang Word", "Ảnh sang Word (OCR)"], horizontal=True)
    
    if mode == "PDF sang Word":
        file_upload = st.file_uploader("Chọn file PDF:", type="pdf")
        is_ocr = st.toggle("Chế độ Scan (OCR) - Dùng cho sách quét", value=False)
        pages = st.text_input("Khoảng trang cần lấy (VD: 1-20):", "")
    else:
        file_upload = st.file_uploader("Chọn các ảnh:", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="main-box" style="min-height:350px;">', unsafe_allow_html=True)
    st.markdown("### ⚡ TRẠNG THÁI & KẾT QUẢ")
    
    if file_upload:
        if st.button("BẮT ĐẦU TRÍCH XUẤT"):
            out_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
            try:
                # Phân tích trang
                p_start, p_end = 0, None
                if mode == "PDF sang Word" and pages.strip():
                    if "-" in pages:
                        s, e = map(int, pages.split("-"))
                        p_start, p_end = s-1, e
                    else:
                        p_start, p_end = int(pages)-1, int(pages)

                with st.spinner("Đang xử lý dữ liệu..."):
                    if mode == "PDF sang Word":
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                            tmp_pdf.write(file_upload.getvalue())
                            tmp_pdf_p = tmp_pdf.name
                        
                        if not is_ocr:
                            cv = Converter(tmp_pdf_p)
                            cv.convert(out_path, start=p_start, end=p_end)
                            cv.close()
                        else:
                            # Tối ưu RAM cho file 50MB: Giảm DPI xuống 150
                            imgs = convert_from_path(tmp_pdf_p, dpi=150, first_page=p_start+1, last_page=p_end)
                            doc = Document()
                            progress_bar = st.progress(0)
                            for i, img in enumerate(imgs):
                                txt = pytesseract.image_to_string(img, lang='vie')
                                doc.add_paragraph(f"--- TRANG {p_start + i + 1} ---")
                                doc.add_paragraph(txt)
                                progress_bar.progress((i + 1) / len(imgs))
                                del img # Giải phóng RAM ngay lập tức
                            doc.save(out_path)
                        os.remove(tmp_pdf_p)
                    else:
                        doc = Document()
                        for img_file in file_upload:
                            txt = pytesseract.image_to_string(Image.open(img_file), lang='vie')
                            doc.add_paragraph(txt)
                        doc.save(out_path)

                st.success("✅ Thành công! Hãy tải file bên dưới:")
                with open(out_path, "rb") as f:
                    st.download_button("📥 TẢI FILE WORD VỀ MÁY", f, file_name=f"Ket_qua.docx")
                os.remove(out_path)
                gc.collect() # Thu dọn rác bộ nhớ
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")
                st.info("Mẹo: Với file nặng > 50MB, hãy chia nhỏ khoảng trang (VD: 1-15) để tránh lỗi 'Oh no'.")
    else:
        st.write("---")
        st.write("👉 Hãy tải file ở bên trái để bắt đầu.")
    st.markdown('</div>', unsafe_allow_html=True)
