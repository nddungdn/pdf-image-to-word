import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pytesseract
from PIL import Image
import os, tempfile, gc, re
from pdf2image import convert_from_path

# Cấu hình giao diện Neon
st.set_page_config(page_title="Học Liệu Số - Định Dạng Chuẩn", layout="wide")

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

# Hàm làm sạch văn bản OCR (Xóa xuống dòng thừa, nối từ bị ngắt)
def clean_text(text):
    text = re.sub(r'(?<=[^\s])\n(?=[a-zà-ỹ])', ' ', text) # Nối các dòng bị ngắt giữa chừng
    text = re.sub(r' +', ' ', text) # Xóa khoảng trắng thừa
    return text.strip()

st.markdown('<h2 style="text-align:center; color:#00d4ff;">✨ TRÍCH XUẤT & TỰ ĐỘNG ĐỊNH DẠNG ĐẸP</h2>', unsafe_allow_html=True)

col1, col2 = st.columns([1, 1.2])

with col1:
    st.markdown('<div class="main-box">', unsafe_allow_html=True)
    st.write("### 📥 CÀI ĐẶT")
    file_upload = st.file_uploader("Tải PDF/Ảnh:", type=["pdf", "jpg", "png", "jpeg"])
    pages = st.text_input("Khoảng trang (VD: 1-10):", "")
    st.info("Hệ thống sẽ tự: Font Times New Roman 13, Giãn dòng 1.5, Căn đều 2 bên, Lề chuẩn A4.")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="main-box" style="min-height:300px;">', unsafe_allow_html=True)
    st.write("### ⚡ KẾT QUẢ")
    
    if file_upload:
        if st.button("🚀 BẮT ĐẦU XỬ LÝ"):
            out_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
            doc = Document()
            
            # 1. Thiết lập lề trang A4 chuẩn (Trên 2cm, Dưới 2cm, Trái 3cm, Phải 1.5cm)
            section = doc.sections[0]
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(3)
            section.right_margin = Cm(1.5)

            try:
                p_start, p_end = 1, None
                if pages.strip() and "-" in pages:
                    s, e = map(int, pages.split("-"))
                    p_start, p_end = s, e

                with st.spinner("Đang quét và định dạng..."):
                    if file_upload.type == "application/pdf":
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                            tmp.write(file_upload.getvalue())
                            tmp_p = tmp.name
                        imgs = convert_from_path(tmp_p, dpi=130, first_page=p_start, last_page=p_end)
                        
                        for i, img in enumerate(imgs):
                            raw_text = pytesseract.image_to_string(img, lang='vie')
                            processed_text = clean_text(raw_text)
                            
                            # 2. Tạo đoạn văn và áp dụng định dạng chuẩn
                            p = doc.add_paragraph()
                            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # Căn đều 2 bên
                            run = p.add_run(processed_text)
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(13) # Cỡ chữ 13
                            p.paragraph_format.line_spacing = 1.5 # Giãn dòng 1.5
                            p.paragraph_format.space_after = Pt(6) # Khoảng cách sau đoạn
                            
                            doc.add_page_break()
                            del img
                        os.remove(tmp_p)
                    else:
                        txt = clean_text(pytesseract.image_to_string(Image.open(file_upload), lang='vie'))
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        run = p.add_run(txt)
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(13)
                        p.paragraph_format.line_spacing = 1.5

                    doc.save(out_path)
                    st.success("✅ Đã hoàn tất định dạng chuẩn!")
                    with open(out_path, "rb") as f:
                        st.download_button("📥 TẢI FILE WORD ĐÃ ĐẸP", f, file_name="Tai_lieu_chuan.docx")
                
                os.remove(out_path)
                gc.collect()
            except Exception as e:
                st.error(f"Lỗi: {str(e)}")
    else:
        st.write("---")
        st.write("Sẵn sàng xử lý file của bạn.")
    st.markdown('</div>', unsafe_allow_html=True)
