import streamlit as st
from pdf2docx import Converter
from docx import Document
import pytesseract
from PIL import Image
import os
import tempfile
from pdf2image import convert_from_path

st.set_page_config(page_title="Document Pro - Chuyên xử lý file Scan", page_icon="📄", layout="wide")

# Sidebar hướng dẫn
with st.sidebar:
    st.title("💡 Lưu ý quan trọng")
    st.warning("""
    **Nếu file Word bị biến thành ảnh:**
    Hãy gạt nút **'Chế độ Scan (OCR)'** bên dưới phần tải file. 
    Hệ thống sẽ dùng AI để 'đọc' chữ từ ảnh cho bạn.
    """)
    st.write("---")
    st.write("📧 admin@thcslethanhton.edu.vn")

st.title("🚀 Chuyển đổi PDF & Ảnh sang Word (Bản Pro)")

tab1, tab2 = st.tabs(["📄 Chuyển PDF sang Word", "🖼️ Chuyển Ảnh sang Word"])

with tab1:
    pdf_file = st.file_uploader("Tải file PDF của bạn", type="pdf")
    
    # Nút gạt quan trọng nhất để xử lý file của bạn
    is_scan = st.toggle("Đây là file PDF dạng ảnh quét (Scan) - Cần dùng OCR để lấy chữ", value=False)
    
    if pdf_file:
        if st.button("🚀 Bắt đầu xử lý PDF"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(pdf_file.getvalue())
                tmp_path = tmp.name
            
            output_docx = tmp_path.replace(".pdf", ".docx")

            try:
                if not is_scan:
                    # Chế độ thông thường: Giữ bố cục tốt cho file PDF có văn bản thật
                    with st.spinner("Đang tái tạo bố cục..."):
                        cv = Converter(tmp_path)
                        cv.convert(output_docx)
                        cv.close()
                else:
                    # Chế độ Scan (OCR): Dành cho file SGK 50MB của bạn
                    with st.spinner("Đang dùng OCR để đọc từng trang sách (có thể mất vài phút)..."):
                        images = convert_from_path(tmp_path)
                        doc = Document()
                        for i, image in enumerate(images):
                            text = pytesseract.image_to_string(image, lang='vie')
                            doc.add_paragraph(f"--- TRANG {i+1} ---")
                            doc.add_paragraph(text)
                            doc.add_page_break()
                        doc.save(output_docx)

                st.success("✅ Đã xử lý xong!")
                with open(output_docx, "rb") as f:
                    st.download_button("📥 Tải file Word về máy", f, file_name=pdf_file.name.replace(".pdf", ".docx"))
            
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")
            finally:
                if os.path.exists(tmp_path): os.remove(tmp_path)

# --- TAB 2 giữ nguyên logic cũ nhưng tối ưu giao diện ---
with tab2:
    img_files = st.file_uploader("Chọn ảnh", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    if img_files and st.button("🔍 Quét chữ từ ảnh"):
        doc = Document()
        for img_file in img_files:
            text = pytesseract.image_to_string(Image.open(img_file), lang='vie')
            doc.add_paragraph(text)
            doc.add_page_break()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            with open(tmp.name, "rb") as f:
                st.download_button("📥 Tải file Word kết quả", f, file_name="Ket_qua_OCR.docx")
