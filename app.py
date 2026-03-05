import streamlit as st
from docx import Document
import pytesseract
from PIL import Image
import os, tempfile, gc
from pdf2image import convert_from_path

# Cấu hình giao diện Neon tối giản
st.set_page_config(page_title="Học Liệu Số - Trích Xuất Văn Bản", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #000b14; color: #00d4ff; }
    .main-box {
        border: 1px solid #00d4ff; border-radius: 10px;
        padding: 15px; background-color: rgba(0, 20, 40, 0.9);
        box-shadow: 0 0 10px #00d4ff;
    }
    .stButton>button { width: 100%; border: 1px solid #00d4ff; background: transparent; color: #00d4ff; font-weight: bold; }
    .stButton>button:hover { background: #00d4ff; color: #000; box-shadow: 0 0 20px #00d4ff; }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<h2 style="text-align:center; color:#00d4ff;">📝 HỆ THỐNG TRÍCH XUẤT VĂN BẢN (KHÔNG LẤY ẢNH)</h2>', unsafe_allow_html=True)

col1, col2 = st.columns([1, 1.2])

with col1:
    st.markdown('<div class="main-box">', unsafe_allow_html=True)
    st.write("### 📥 CÀI ĐẶT")
    file_upload = st.file_uploader("Tải lên PDF hoặc Ảnh:", type=["pdf", "jpg", "png", "jpeg"], accept_multiple_files=False)
    pages = st.text_input("Khoảng trang (VD: 1-10):", help="Nên chia nhỏ nếu file > 50 trang")
    st.info("Chế độ này sẽ loại bỏ mọi hình ảnh, chỉ giữ lại chữ viết để bạn soạn bài.")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="main-box" style="min-height:300px;">', unsafe_allow_html=True)
    st.write("### ⚡ KẾT QUẢ")
    
    if file_upload:
        if st.button("BẮT ĐẦU LẤY CHỮ"):
            out_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
            doc = Document()
            
            try:
                # Xử lý khoảng trang
                p_start, p_end = 1, None
                if pages.strip() and "-" in pages:
                    s, e = map(int, pages.split("-"))
                    p_start, p_end = s, e

                with st.spinner("Đang quét chữ Tiếng Việt..."):
                    if file_upload.type == "application/pdf":
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
                            tmp_pdf.write(file_upload.getvalue())
                            tmp_pdf_p = tmp_pdf.name
                        
                        # Sử dụng Poppler để chụp ảnh từng trang với DPI thấp (tiết kiệm RAM)
                        imgs = convert_from_path(tmp_pdf_p, dpi=130, first_page=p_start, last_page=p_end)
                        bar = st.progress(0)
                        
                        for i, img in enumerate(imgs):
                            # Quét OCR Tiếng Việt
                            text = pytesseract.image_to_string(img, lang='vie')
                            doc.add_heading(f"Trang {p_start + i}", level=2)
                            doc.add_paragraph(text)
                            bar.progress((i + 1) / len(imgs))
                            del img # Xóa ảnh khỏi RAM ngay
                        os.remove(tmp_pdf_p)
                    else:
                        # Xử lý nếu là file ảnh đơn lẻ
                        text = pytesseract.image_to_string(Image.open(file_upload), lang='vie')
                        doc.add_paragraph(text)
                    
                    doc.save(out_path)
                    st.success("Đã trích xuất xong văn bản!")
                    with open(out_path, "rb") as f:
                        st.download_button("📥 TẢI VĂN BẢN (.DOCX)", f, file_name="Van_ban_trich_xuat.docx")
                
                os.remove(out_path)
                gc.collect() # Dọn dẹp bộ nhớ RAM
            except Exception as e:
                st.error(f"Lỗi hệ thống: {str(e)}")
                st.info("Lưu ý: Đảm bảo 'packages.txt' đã có 'poppler-utils'.")
    else:
        st.write("---")
        st.write("Vui lòng tải file để bắt đầu.")
    st.markdown('</div>', unsafe_allow_html=True)
