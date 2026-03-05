import streamlit as st
from pdf2docx import Converter
from docx import Document
import pytesseract
from PIL import Image
import os
import tempfile

# Cấu hình giao diện
st.set_page_config(page_title="Chuyển đổi tài liệu chuyên nghiệp", layout="wide")

st.title("📄 Tiện ích chuyển đổi PDF & Ảnh sang Word")
st.markdown("Hỗ trợ file lớn trên 100 trang, giữ nguyên bố cục và tối ưu chính tả Tiếng Việt.")

tab1, tab2 = st.tabs(["Chuyển PDF sang Word", "Chuyển Ảnh sang Word"])

# --- XỬ LÝ PDF ---
with tab1:
    pdf_file = st.file_uploader("Kéo thả file PDF vào đây", type="pdf")
    if pdf_file:
        if st.button("🚀 Bắt đầu chuyển PDF"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(pdf_file.getvalue())
                tmp_path = tmp.name
            
            output_path = tmp_path.replace(".pdf", ".docx")
            with st.spinner("Đang phân tích bố cục và chuyển đổi..."):
                cv = Converter(tmp_path)
                cv.convert(output_path)
                cv.close()
            
            with open(output_path, "rb") as f:
                st.download_button("📥 Tải file Word về máy", f, file_name=pdf_file.name.replace(".pdf", ".docx"))
            os.remove(tmp_path)

# --- XỬ LÝ ẢNH (OCR) ---
with tab2:
    img_files = st.file_uploader("Chọn một hoặc nhiều ảnh", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    if img_files:
        if st.button("🔍 Quét OCR Tiếng Việt"):
            doc = Document()
            with st.spinner("Đang nhận diện chữ từ ảnh..."):
                for img_file in img_files:
                    img = Image.open(img_file)
                    # Nhận diện chữ Tiếng Việt
                    text = pytesseract.image_to_string(img, lang='vie')
                    doc.add_paragraph(text)
                    doc.add_page_break()
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    doc.save(tmp.name)
                    with open(tmp.name, "rb") as f:
                        st.download_button("📥 Tải file Word đã quét", f, file_name="Chuyển_đổi_ảnh.docx")
